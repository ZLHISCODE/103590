VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObjectInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'类说明：
'适用范围：现在仅限于表以及列检查。其余类型暂时尚未支持
'使用方法：
'               创建一个临时类用来获取该类的新实例
'               Set objTmp=New clsObjectInfo --用来获取该类的新实例
'               Set objParent=objTmp.GetObject(....)--获取需要实例的父级实例
'               Set Objcurrent=objTmp.GetObject(....,objParent)--获取当前实例，并设置实例关系
'               使用对象检查时，检查并修复子级会自动检查修复父级。

Public Enum ObjectType
    OT_Table = 0
    OT_Column = 1
    OT_Constraint = 2
    OT_Index = 3
    OT_Sequence = 4
End Enum
Private mstrObjectName As String
Private motObjectType As ObjectType
Private mobjParent As clsObjectInfo
Private mstrOwner As String
Private mstrAdjustSQL As String
Private mblnChecked As Boolean '对象已经检查成功
Private mcnOwner As ADODB.Connection '所有者的连接

'ObjectName:对象名称
Public Property Get ObjectName() As String
    ObjectName = mstrObjectName
End Property
Public Property Let ObjectName(ByVal strObjectName As String)
    mstrObjectName = strObjectName
End Property
'ObjectType:对象类型
Public Property Get ObjectType() As ObjectType
    ObjectType = motObjectType
End Property

Public Property Let ObjectType(ByVal otObjectType As ObjectType)
    motObjectType = otObjectType
End Property
'ObjectParent:父对象
Public Property Get ObjectParent() As clsObjectInfo
    Set ObjectParent = mobjParent
End Property

Public Property Set ObjectParent(ByVal objParent As clsObjectInfo)
    Set mobjParent = objParent
End Property
'ObjectOwner:对象所属用户
Public Property Get ObjectOwner() As String
    ObjectOwner = mstrOwner
End Property

Public Property Let ObjectOwner(ByVal strOwner As String)
    mstrOwner = strOwner
End Property
'AdjustSQL:修正SQL
Public Property Get AdjustSQL() As String
    AdjustSQL = mstrAdjustSQL
End Property

Public Property Let AdjustSQL(ByVal strAdjustSQL As String)
    mstrAdjustSQL = strAdjustSQL
End Property
'Checked:是否已经通过检查
Public Property Get Checked() As Boolean
    Checked = mblnChecked
End Property

Public Property Let Checked(ByVal blnChecked As Boolean)
    mblnChecked = blnChecked
End Property
'OwnerConn:所有者连接
Public Property Get OwnerConn() As ADODB.Connection
    Set OwnerConn = mcnOwner
End Property

Public Property Set OwnerConn(ByVal cnOwner As ADODB.Connection)
    Set mcnOwner = cnOwner
End Property

Public Function GetObject(ByVal strObjectName As String, ByVal otObjectType As ObjectType, Optional ByVal strAdjustSQL As String, _
                                        Optional ByVal strObjectOwner As String = "ZLTOOLS", Optional ByVal objParent As clsObjectInfo, Optional ByVal cnOwner As ADODB.Connection) As clsObjectInfo
'功能:创建对象，该方法使用具体查看《类说明》在该类开头
'参数：strObjectName=对象名
'         otObjectType=对象类型
'         strAdjustSQL=修正SQL
'         strObjectOwner=对象所属用户
'         objParent=当前对象的父级对象
' 返回：当前对象
    Dim objCurent As clsObjectInfo
    Set objCurent = New clsObjectInfo
    objCurent.ObjectName = strObjectName
    objCurent.ObjectType = otObjectType
    objCurent.AdjustSQL = strAdjustSQL
    objCurent.ObjectOwner = strObjectOwner
    Set objCurent.ObjectParent = objParent
    Set objCurent.OwnerConn = cnOwner
    Set GetObject = objCurent
End Function

Public Function ObjectCheck(Optional ByVal blnMsg As Boolean = True) As Boolean
'功能：对当前对象进行检查以及修复（会自动对当期对象的父级进行检查修复）
'参数：blnMsg=是否进行信息提示
'返回：是否检查成功或者正常修正
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strMsg As String
    Dim arrAdjustSQL As Variant, strOtherSQL As String
    Dim strTableName As String
    Dim blnOk As Boolean, i As Integer
    Dim cnTmp As ADODB.Connection
    
    On Error GoTo errH
    '存在父级，则检查父级，父级检查失败，则退出
    If Not Me.ObjectParent Is Nothing Then
        If Not Me.ObjectParent.Checked Then '父对象未通过检查
            If Not Me.ObjectParent.ObjectCheck(blnMsg) Then
                Exit Function
            Else
                Me.ObjectParent.Checked = True
            End If
        End If
    End If
    If Not Me.Checked Then '是否已经通过检查
        '进行对象检查
        Select Case Me.ObjectType
            Case OT_Table
                If Me.OwnerConn Is Nothing Then
                    strSQL = "Select 1 FROM ALL_TABLES WHERE TABLE_NAME=[1] AND OWNER=[2]"
                    strMsg = "表""" & Me.ObjectOwner & "." & Me.ObjectName & """不存在，请手工创建该表" & IIf(UCase(Me.ObjectOwner) = "ZLTOOLS", "，同时创建同义词并授权。", "。")
                Else
                    strSQL = "Select 1 FROM User_TABLES WHERE TABLE_NAME=[1]"
                End If
                If UCase(Me.ObjectOwner) = "ZLTOOLS" Then
                    strOtherSQL = "create public synonym " & Me.ObjectName & "  for  ZLTOOLS." & Me.ObjectName
                    strOtherSQL = strOtherSQL & "|Grant Select On ZLTOOLS." & Me.ObjectName & " to PuBLIC"
                    strOtherSQL = strOtherSQL & "|Grant Select,Insert,Update,Delete On ZLTOOLS." & Me.ObjectName & " to " & gstrUserName & "With Grant Option"
                End If
                strTableName = Me.ObjectName
            Case OT_Column
                If Me.OwnerConn Is Nothing Then
                    strSQL = "Select 1 From All_Tab_Columns Where Owner =[2] And Table_Name = [1] And Column_Name = [3]"
                Else
                    strSQL = "Select 1 From User_Tab_Columns Where Table_Name = [1] And Column_Name = [3]"
                End If
                strMsg = "表""" & Me.ObjectOwner & "." & Me.ObjectParent.ObjectName & """的列""" & Me.ObjectName & """ 不存在，请手工添加该列。"
                strTableName = Me.ObjectParent.ObjectName
        End Select
        Set rsTmp = gclsBase.OpenSQLRecord(IIf(Me.OwnerConn Is Nothing, gcnOracle, Me.OwnerConn), strSQL, gstrSysName, UCase(strTableName), UCase(Me.ObjectOwner), UCase(Me.ObjectName))
        blnOk = Not rsTmp.EOF
        If Not blnOk Then
            If Me.AdjustSQL = "" Then
                If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
            Else
                If UCase(Me.ObjectOwner) = "ZLTOOLS" Then
                    Set cnTmp = GetConnection("ZLTOOLS")
                    If cnTmp Is Nothing Then
                        Set cnTmp = gcnOracle
                    End If
                ElseIf mcnOwner Is Nothing Then
                    Set cnTmp = gcnOracle
                Else
                    Set cnTmp = mcnOwner
                End If
                arrAdjustSQL = Split(Me.AdjustSQL, "|")
                On Error Resume Next
                For i = LBound(arrAdjustSQL) To UBound(arrAdjustSQL)
                    cnTmp.Execute arrAdjustSQL(i)
                    If err.Number <> 0 Then
                        err.Clear
                        '修复表时，表创建出错，则自动退出
                        '列修复失败，则自动退出
                        If i = LBound(arrAdjustSQL) Then
                            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
                            Exit For
                        End If
                    End If
                Next
                If strOtherSQL <> "" Then
                    arrAdjustSQL = Split(strOtherSQL, "|")
                    For i = LBound(arrAdjustSQL) To UBound(arrAdjustSQL)
                        cnTmp.Execute arrAdjustSQL(i)
                        If err.Number <> 0 Then err.Clear
                    Next
                End If
                On Error GoTo errH
                Set rsTmp = gclsBase.OpenSQLRecord(IIf(Me.OwnerConn Is Nothing, gcnOracle, Me.OwnerConn), strSQL, gstrSysName, UCase(strTableName), UCase(Me.ObjectOwner), UCase(Me.ObjectName))
                blnOk = Not rsTmp.EOF
            End If
        End If
        ObjectCheck = blnOk
        Me.Checked = blnOk
    Else
        ObjectCheck = Me.Checked
    End If
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function

