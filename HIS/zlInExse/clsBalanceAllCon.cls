VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBalanceAllCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'��������ֵ�ľֲ�����
Private mvarstrAllTime As String '�ֲ����ƣ���������δ����סԺ����
Private mvarstrUnAuditTime As String '�ֲ����ƣ���������δ���סԺ����
Private mvarstrAllDeptIDs As String '�ֲ����ƣ���������δ���ʿ���
Private mvarstrAllItem As String '�ֲ����ƣ���������δ���վݷ�Ŀ
Private mvarstrAllClass As String '�ֲ����ƣ���������δ���������
Private mvarstrAllChargeType As String '�ֲ����ƣ���������δ����շ����
Private mvarstrAllOwnerFeeType As String '�ֲ����ƣ���������δ����Էѵ��շ����
Private mvarstrAllDiag As String '�ֲ����ƣ��������
Private mvarstrAllBabys As String '�ֲ����ƣ�����Ӥ��
Private mvarMinDate As Date '�ֲ�����
Private mvarMaxDate As Date '�ֲ�����
Private mvarrsALLTime As ADODB.Recordset '�ֲ�����
Private mvarstrAllFullTims As String '�ֲ�����
Public Property Let strAllFullTims(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.strAllFullTims = 5
    mvarstrAllFullTims = vData
End Property


Public Property Get strAllFullTims() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.strAllFullTims
    strAllFullTims = mvarstrAllFullTims
End Property




Public Property Set rsAllTime(ByVal vData As ADODB.Recordset)
    '������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
    'Syntax: Set x.rsALLTime = Form1
    Set mvarrsALLTime = vData
End Property


Public Property Get rsAllTime() As ADODB.Recordset
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.rsALLTime
    Set rsAllTime = mvarrsALLTime
End Property
Public Property Let MaxDate(ByVal vData As Date)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.MaxDate = 5
    mvarMaxDate = vData
End Property


Public Property Get MaxDate() As Date
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.MaxDate
    MaxDate = mvarMaxDate
End Property



Public Property Let MinDate(ByVal vData As Date)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.MinDate = 5
    mvarMinDate = vData
End Property


Public Property Get MinDate() As Date
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.MinDate
    MinDate = mvarMinDate
End Property



Public Property Let strAllBabys(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.strAllBabys = 5
    mvarstrAllBabys = vData
End Property


Public Property Get strAllBabys() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.strAllBabys
    strAllBabys = mvarstrAllBabys
End Property



Public Property Let strAllOwnerFeeType(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.strAllOwnerFeeType = 5
    mvarstrAllOwnerFeeType = vData
End Property


Public Property Get strAllOwnerFeeType() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.strAllOwnerFeeType
    strAllOwnerFeeType = mvarstrAllOwnerFeeType
End Property



Public Property Let strAllChargeType(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.strAllChargeType = 5
    mvarstrAllChargeType = vData
End Property


Public Property Get strAllChargeType() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.strAllChargeType
    strAllChargeType = mvarstrAllChargeType
End Property


Public Property Let strAllDiag(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.strAllChargeType = 5
    mvarstrAllDiag = vData
End Property


Public Property Get strAllDiag() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.strAllChargeType
    strAllDiag = mvarstrAllDiag
End Property


Public Property Let strAllClass(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.strAllClass = 5
    mvarstrAllClass = vData
End Property


Public Property Get strAllClass() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.strAllClass
    strAllClass = mvarstrAllClass
End Property



Public Property Let strAllItem(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.strAllItem = 5
    mvarstrAllItem = vData
End Property


Public Property Get strAllItem() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.strAllItem
    strAllItem = mvarstrAllItem
End Property



Public Property Let strAllDeptIDs(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.strAllDeptIDs = 5
    mvarstrAllDeptIDs = vData
End Property


Public Property Get strAllDeptIDs() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.strAllDeptIDs
    strAllDeptIDs = mvarstrAllDeptIDs
End Property



Public Property Let strUnAuditTime(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.strUnAuditTime = 5
    mvarstrUnAuditTime = vData
End Property


Public Property Get strUnAuditTime() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.strUnAuditTime
    strUnAuditTime = mvarstrUnAuditTime
End Property



Public Property Let strAllTime(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.strAllTime = 5
    mvarstrAllTime = vData
End Property


Public Property Get strAllTime() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.strAllTime
    strAllTime = mvarstrAllTime
End Property

Public Function zlGetTimeRecordFromTimeString(ByVal lng����ID As Long, ByVal strAllTimes As String, _
    ByRef rsTimes As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����סԺ����������ȡÿ����Ϣ��Ϣ
    '���:
    '����:rsTimes-סԺ��Ϣ��(����ID,��ҳID,����,��������)
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2017-11-12 14:37:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    Set rsTimes = Nothing
    
    strSQL = "" & _
    "   Select A.����ID,A.��ҳID,Nvl(A.����,0) As ����,B.���� as ��������" & _
    "   From ������ҳ A,������� B " & _
    "   Where A.����=B.���(+) And A.����ID = [1]  " & _
    "       And A.��ҳID In (Select Column_Value From Table(f_str2list([2]))) " & _
    "   Order By A.����ID,A.��ҳID Desc"
    
    Set rsTimes = zlDatabase.OpenSQLRecord(strSQL, "��ȡָ��������סԺ��Ϣ", lng����ID, strAllTimes)
    
    zlGetTimeRecordFromTimeString = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetNumsFromZyNo(ByVal strסԺ�� As String, Optional ByRef lng����ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����סԺ������ȡסԺ����
    '���:lngסԺ��
    '����:lng����ID-���ز���ID
    '����:סԺ����,����ö��ŷ���
    '����:���˺�
    '����:2017-11-12 10:40:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strTimeNums As String
     
    lng����ID = 0
    strSQL = "Select ����ID,��ҳID From ������ҳ Where סԺ�� = [1] And Nvl(��������,0) = 0 Order By ��ҳID DESC"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����סԺ������ȡ���漰��סԺ����", strסԺ��)
    strTimeNums = ""
    Do While Not rsTemp.EOF
        lng����ID = Val(Nvl(rsTemp!����ID))
        strTimeNums = strTimeNums & "," & Val(Nvl(rsTemp!��ҳID))
        rsTemp.MoveNext
    Loop
    If strTimeNums <> "" Then strTimeNums = Mid(strTimeNums, 2)
    zlGetNumsFromZyNo = strTimeNums
End Function
Public Function GetPatiFeeDateRang(ByVal lng����ID As Long, str��ҳIds As String, ByRef strStartDate As String, ByRef strEndDate As String, ByRef bln�Ǽ�ʱ�� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���˵ķ���ʱ�䷶Χ
    '���:lng����ID
    '     str��ҳIDs-��ҳID,����ö��ŷָ�
    '����:strStartDate-����ʱ��
    '     strEndDate-����ʱ��
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2017-11-26 10:49:03
    '˵��:����ȡ����������:
    '    һ����ʼʱ��
    '       1.�����Ժ����С�ڷ���ʱ�䣬��ȡ��Ժʱ��
    '       2.�����Ժ���ڴ��ڷ���ʱ�䣬��ȡ����ʱ��
    '    ��������ʱ��
    '       1.�����Ժ����С�ڷ���ʱ�䣬��ȡ����ʱ��
    '       2.�����Ժ���ڴ��ڷ���ʱ�䣬��ȡ��Ժʱ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFeeMinDate As String, strFeeMaxDate As String
    Dim str��Ժ���� As String, str��Ժ���� As String
    Dim dtDate As Date, dtTemp As Date
    
    On Error GoTo errHandle
    strStartDate = "": strEndDate = ""
    
    'ȡ���õ���С�����ʱ��
    Call GetPatiFeeMinAndMaxDate(lng����ID, str��ҳIds, strFeeMinDate, strFeeMaxDate, bln�Ǽ�ʱ��)
    Call GetPatiInOutDate(lng����ID, str��ҳIds, str��Ժ����, str��Ժ����)
    
    
    If str��Ժ���� <> "" Then
        dtDate = CDate(str��Ժ����)
        If strFeeMinDate <> "" Then
            dtTemp = CDate(strFeeMinDate)
            If dtDate > dtTemp Then dtDate = dtTemp
        End If
    Else
        dtDate = zlDatabase.Currentdate
    End If
    strStartDate = Format(dtDate, "yyyy-mm-dd HH:MM:SS")
    
    If str��Ժ���� <> "" Then
        dtDate = CDate(str��Ժ����)
        If strFeeMaxDate <> "" Then
            dtTemp = CDate(strFeeMaxDate)
            If dtDate < dtTemp Then dtDate = dtTemp
        End If
        strEndDate = Format(dtDate, "yyyy-mm-dd HH:MM:SS")
    Else
        'δ��Ժ
        If strFeeMaxDate <> "" Then
            dtDate = CDate(strFeeMaxDate)
        Else
            dtDate = zlDatabase.Currentdate
        End If
        strEndDate = Format(dtDate, "yyyy-mm-dd HH:MM:SS")
    End If
    If strStartDate > strEndDate Then
        strEndDate = strStartDate
    End If
    GetPatiFeeDateRang = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetPatiInOutDate(lng����ID As Long, str��ҳIds As String, ByRef str��Ժʱ�� As String, ByRef str��Ժ���� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����ĳ��סԺ����Ժ���Ժʱ��
    '���:
    '����:str��Ժʱ��-��ʽ:yyyy-mm-dd HH:MM:SS����ҳIDΪ���ʱ,������С����Ժʱ��
    '     str��Ժ����-��ʽ:yyyy-mm-dd HH:MM:SS����ҳIDΪ���ʱ,����������Ժʱ��,δ��Ժʱ����Ժ����Ϊ""
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2017-11-26 11:09:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strWhere As String, lng��ҳID As Long
    
    On Error GoTo errHandle
    
    If InStr(str��ҳIds, ",") > 0 Then
        strWhere = " And ��ҳID in (Select Column_Value From Table(f_num2list([2]))) "
    Else
        strWhere = " And ��ҳID= [3] ": lng��ҳID = Val(str��ҳIds)
    End If
    
    strSQL = "Select min(��Ժ����) as ��Ժ����,max(��Ժ����) as ��Ժ���� From ������ҳ Where ����ID=[1]  " & strWhere
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���˵����Ժ����", lng����ID, str��ҳIds, lng��ҳID)
    
    str��Ժʱ�� = Format(rsTmp!��Ժ����, "yyyy-mm-dd HH:MM:SS")
    str��Ժ���� = Format(rsTmp!��Ժ����, "yyyy-mm-dd HH:MM:SS")
 
    GetPatiInOutDate = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetPatiFeeMinAndMaxDate(ByVal lng����ID As Long, ByVal str��ҳIds As String, ByRef strMinDate As String, ByRef strMaxDate As String, Optional bln�Ǽ�ʱ�� As Boolean) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����˵ĵǼ�ʱ��
    '���:lng����ID
    '     str��ҳIDs-��ҳID,����ö��ŷָ�
    '����:strMinDate-��ʽ:yyyy-mm-dd HH:MM:SS
    '     strMaxDate-��ʽ:yyyy-mm-dd HH:MM:SS
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2013-10-22 17:16:47
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strFields As String, strWhere As String, lng��ҳID As Long
    
    On Error GoTo errHandle
    strFields = IIf(bln�Ǽ�ʱ��, "�Ǽ�ʱ��", "����ʱ��")
    strFields = "Min(" & strFields & ") as ��Сʱ��,Max(" & strFields & ") as ���ʱ�� "
    If InStr(str��ҳIds, ",") > 0 Then
        strWhere = " And ��ҳID in (Select Column_Value From Table(f_num2list([2]))) "
    Else
        strWhere = " And ��ҳID= [3] ": lng��ҳID = Val(str��ҳIds)
    End If
    strSQL = " Select " & strFields & " From סԺ���ü�¼ where mod(��¼����,10) in (2,3) And ����ID=[1] " & strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���õ���С�����ʱ��", lng����ID, str��ҳIds, lng��ҳID)
    
    strMinDate = Format(rsTemp!��Сʱ��, "yyyy-mm-dd HH:MM:SS")
    strMaxDate = Format(rsTemp!���ʱ��, "yyyy-mm-dd HH:MM:SS")
    GetPatiFeeMinAndMaxDate = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


