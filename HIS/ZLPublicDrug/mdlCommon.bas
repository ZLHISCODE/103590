Attribute VB_Name = "mdlCommon"
Option Explicit

Public gstrSysName As String                'ϵͳ����
Public gstrUnitName As String               '�û���λ����
Public gstrProductName As String    '��Ʒ����
Public gstrSQL As String
Public glngSys As Long
Public glngMainModule As Long '�����ߵ�ģ���
Public gstrMainPrivs As String '�����ߵ����Ȩ��
Public gcnOracle As ADODB.Connection
Public grsStockCheck As ADODB.Recordset      '�����
Public gstrDBUser As String '������

'����������
Public gobjComlib As Object
Public gobjCommFun As Object
Public gobjControl As Object
Public gobjDatabase As Object
Public gstrNodeNo As String 'վ����

'�ӿ�Ҫʹ�õ���ϵͳ����
Public Type Type_SysParms
    P9_���ý���λ�� As Integer
    P150_ҩƷ���������㷨 As Integer
    P157_���õ��۱���λ�� As Integer
End Type
Public gtype_UserSysParms As Type_SysParms     'ϵͳ����

Public Enum StockCheck
    ����� = 0
    �������� = 1
    �����ֹ = 2
End Enum

'�û���Ϣ------------------------
Public Type TYPE_USER_INFO
    �û�ID As Long
    �û����� As String
    �û����� As String
    �û����� As String
    ����ID As Long
    ���ű��� As String
    �������� As String
    strMaterial As String
End Type
Public UserInfo As TYPE_USER_INFO

Public Function zlGetComLib() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����������ض���
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-05-15 15:34:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not gobjComlib Is Nothing Then zlGetComLib = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set gobjComlib = GetObject("", "zl9Comlib.clsComlib")
    Set gobjCommFun = GetObject("", "zl9Comlib.clsCommfun")
    Set gobjControl = GetObject("", "zl9Comlib.clsControl")
    Set gobjDatabase = GetObject("", "zl9Comlib.clsDatabase")
    gstrNodeNo = ""
    If Not gobjComlib Is Nothing Then gstrNodeNo = gobjComlib.gstrNodeNo
    Err = 0: On Error GoTo 0
    If Not gobjComlib Is Nothing Then zlGetComLib = True: Exit Function
    Err = 0: On Error Resume Next
    Set gobjComlib = CreateObject("zl9Comlib.clsComlib")
    Call gobjComlib.InitCommon(gcnOracle)
    Set gobjCommFun = gobjComlib.zlCommFun
    Set gobjControl = gobjComlib.zlControl
    Set gobjDatabase = gobjComlib.gobjDatabase
    If Not gobjComlib Is Nothing Then gstrNodeNo = gobjComlib.gstrNodeNo
    Err = 0: On Error GoTo 0
End Function
 
Public Sub GetSysParms()
    'ȡϵͳ����ֵ
    
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    
    gstrSQL = "Select ������, ����ֵ, ȱʡֵ From Zlparameters Where ϵͳ = 100 And Nvl(˽��, 0) = 0 And ģ�� Is Null Order By ������ "
    Set rsTemp = gobjDatabase.OpenSQLRecord(gstrSQL, "GetSysParms")
    
    If rsTemp.RecordCount > 0 Then
        rsTemp.Filter = "������=9"
        If Not rsTemp.EOF Then gtype_UserSysParms.P9_���ý���λ�� = Val(NVL(rsTemp!����ֵ, rsTemp!ȱʡֵ))
        
        rsTemp.Filter = "������=150"
        If Not rsTemp.EOF Then gtype_UserSysParms.P150_ҩƷ���������㷨 = Val(NVL(rsTemp!����ֵ, rsTemp!ȱʡֵ))
        
        rsTemp.Filter = "������=157"
        If Not rsTemp.EOF Then gtype_UserSysParms.P157_���õ��۱���λ�� = Val(NVL(rsTemp!����ֵ, rsTemp!ȱʡֵ))
    End If
    
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
    '���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Sub GetStockCheckRule()
    'ȡ��������
    
    gstrSQL = "Select �ⷿid, ��鷽ʽ From ҩƷ������ "
    Set grsStockCheck = gobjDatabase.OpenSQLRecord(gstrSQL, "GetStockCheckRule")
    
End Sub

Public Sub GetUserInfo()
    Dim rsUser As ADODB.Recordset
    
    Set rsUser = gobjDatabase.GetUserInfo
    
    With rsUser
        If Not .EOF Then
            UserInfo.�û�ID = !Id
            UserInfo.�û����� = !���
            UserInfo.�û����� = IIf(IsNull(!����), "", !����)
            UserInfo.�û����� = IIf(IsNull(!����), "", !����)
            UserInfo.����ID = !����ID
            UserInfo.���ű��� = !������
            UserInfo.�������� = !������
        Else
            UserInfo.�û�ID = 0
            UserInfo.�û����� = ""
            UserInfo.�û����� = ""
            UserInfo.�û����� = ""
            UserInfo.����ID = 0
            UserInfo.���ű��� = ""
            UserInfo.�������� = ""
        End If
    End With
End Sub
