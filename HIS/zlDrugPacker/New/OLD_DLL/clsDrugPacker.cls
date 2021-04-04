VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrugPacker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub Class_Initialize()
    Set gcolConn = New Collection
    Set gcolDevice = New Collection
End Sub

Private Sub Class_Terminate()
    Dim frmItem As Form
    
    On Error Resume Next
    For Each frmItem In Forms
        Unload frmItem
    Next
    
    Set gcolDevice = Nothing
    Set gcolConn = Nothing
    Set frmItem = Nothing
    Set gobjComLib = Nothing
End Sub

Public Function InitCommon(ByVal cnOracle As ADODB.Connection, ByVal frmOwner As Object, _
        ByVal lngSys As Long, ByVal lngModule As Long, Optional ByVal blnInitLink As Boolean) As Boolean
'���ܣ���ʼ������
'������
'  cnOracle��HIS��DB���Ӷ���
'  frmOwner�������������
'  lngSys������ϵͳ��
'  lngmodule������ģ���
'  blnInitLink���Ƿ��ʼ���Զ����豸�����ӡ�ͨ��������ϴ�������ҵ������ΪTrue��

    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    Dim strKey As String
    Dim frmTmp As frmSetting
    Dim blnCalled As Boolean
    Dim blnMenuItem As Boolean
    
    On Error GoTo errHandle
    
    Set gobjConn = cnOracle
    Set gfrmOwner = frmOwner
    glngSys = lngSys
    glngModule = lngModule
    
    '����zl9Comlib����
    On Error Resume Next
    'Set gobjComLib = New clsComLib
    Set gobjComLib = CreateObject("zl9ComLib.clsComLib")
    If Err <> 0 Or gobjComLib Is Nothing Then
        gstrMessage = "����zl9ComLib����ʧ�ܣ�" & vbNewLine & Err.Description
        Err.Clear
        Exit Function
    End If
    On Error GoTo errHandle
    
    'ע��ҽԺ����
    gstrRegHospital = gobjComLib.zlRegInfo("��λ����", , -1)
    
    '��ȡ�û���Ϣ
    Call GetUserInfo
    
    '��ȡ�ӿڵ�Ȩ��
    '?????
    gstrPrivs = ";�豸����;�ϴ���������;ҩƷ��Ϣ�ϴ�;ҩƷ����ϴ�;"
    
    '��ȡ�豸����
    Call GetDeviceParam
    
    'Ƕ��˵�
    If Not gfrmOwner.mnuDrugPackerItems Is Nothing Then
        If InStr(gstrPrivs, ";ҩƷ��Ϣ�ϴ�;") > 0 Then
            Call SetMenuItem(enuMenuNo.ҩƷ��Ϣ)
            blnMenuItem = True
        End If
        If InStr(gstrPrivs, ";ҩƷ����ϴ�;") > 0 Then
            Call SetMenuItem(enuMenuNo.ҩƷ���)
            blnMenuItem = True
        End If
        If InStr(gstrPrivs, ";�豸����;") > 0 Then
            Call SetMenuItem(enuMenuNo.�豸����)
            blnMenuItem = True
        End If
        If InStr(gstrPrivs, ";�ϴ���������;") > 0 Then
            Call SetMenuItem(enuMenuNo.�ϴ�����)
            blnMenuItem = True
        End If
        '��ʾ�˵�
        If blnMenuItem Then Call ShowMenuItem
    End If
    
    If blnInitLink = False Then
        InitCommon = True
        Exit Function
    End If
    
    
    
    
makInit:
    '�����Զ���ҩ��������
    On Error GoTo errSQL
    gstrSQL = "select ID, ����, ��������, �������� from ҩ���豸���� "
    Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "��ȡ�Զ���ϵͳ������")
    On Error GoTo errHandle
    
    If rsTmp.EOF Then
        rsTmp.Close
        
        If blnCalled = False Then
            
            If MsgBox("��δ���Զ���ϵͳ�������ӣ������Ƿ����ã�", vbInformation + vbYesNo + vbDefaultButton1, GSTR_INTERFACE_NAME) = vbNo Then
                Exit Function
            End If
        
            '���������豸����
            Set frmTmp = New frmSetting
            frmTmp.Show vbModal, frmOwner
            Set frmTmp = Nothing
            
            blnCalled = True
            GoTo makInit
        
        End If
        
    Else
        
        Do While Not rsTmp.EOF
            strKey = rsTmp!����
            
            gcolConn.Add New clsConnect, strKey
            gcolConn(strKey).Name = strKey
            gcolConn(strKey).LinkType = gobjComLib.zlCommFun.NVL(rsTmp!��������, 0)
            
            Select Case gcolConn(strKey).LinkType
                Case enuLinkType.DB
                    With gcolConn(strKey)
                        .DBConnect = New ADODB.Connection
                        On Error Resume Next
                        .DBConnect.Open rsTmp!��������
                        If Err <> 0 Then
                            .Status = False
                            gstrMessage = "��������" & strKey & vbNewLine & _
                                          "���ݣ�" & Err.Description
                        Else
                            .Status = True
                        End If
                        Err.Clear: On Error GoTo 0
                    End With
                Case enuLinkType.WEBServices, enuLinkType.Directory
                    With gcolConn(strKey)
                        .Connect = rsTmp!��������
                        If .Status = False Then
                            gstrMessage = "��������" & strKey & vbNewLine & _
                                          "���ݣ�" & gstrMessage
                        End If
                    End With
            End Select
            rsTmp.MoveNext
        Loop
        
    End If
    rsTmp.Close
    
    InitCommon = True
    Exit Function
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
    Exit Function
    
errSQL:
    If gobjComLib.ErrCenter = 1 Then Resume
End Function

Public Property Get Message() As String
   Message = gstrMessage
End Property

'Public Sub SetInterface(ByVal cnOracle As ADODB.Connection, ByVal frmOwner As Object, ByVal lngSys As Long, ByVal lngModule As Long)
Public Sub SetInterface(ByVal bytType As Byte, Optional ByVal strKey As String)
'���ܣ����๦������
'������
'  bytType��
'    0���ӿڲ�������
'    1���ϴ���������
'    2���豸����

    Dim frmTmp As Form
    
    Select Case bytType
        Case 1
            Set frmTmp = New frmDeviceSetting
            frmTmp.ShowMe 53
        Case 2
            Set frmTmp = New frmDeviceState
            frmTmp.ShowMe 53
        Case Else
            Set frmTmp = New frmSetting
            frmTmp.Show vbModal, gfrmOwner
    End Select
    Set frmTmp = Nothing
    
End Sub

Public Function Upload(ByVal lngModule As Long, ByVal bytType As Byte, Optional ByVal varKey As Variant) As Boolean
'���ܣ�֪ͨ�ӿڰ�ҵ����Ҫ�ϴ�����
'������
'   lngModule��ģ��ţ������ж�ҵ��
'   bytType��
'       0: ҩƷ������Ϣ�ϴ�
'       1: ���ﴦ���ϴ� (��ҩ)
'       2: ���﷢ҩ֪ͨ (��ҩ)
'       3: סԺҩƷҽ���ϴ� (�䡢��ҩ)
'       4: ҩƷ����ϴ�
'   varKey��
'       ��bytType=0ʱ��varKey����д
'       ��bytType=1ʱ��varKey��ʾ������;�ⷿID;NO����
'       ��ʽ��������;�ⷿID;NO[|����;�ⷿID;NO][|...]��
'       ��bytType=2ʱ��ͬbytType=1
'       ��bytType=3ʱ��varKey��ʾҩƷ�շ�ID��
'       ��ʽ����ҩƷ�շ�ID[,ҩƷ�շ�ID][,...]��
'���أ�True�ɹ���Falseʧ��
    
    Dim rsData As ADODB.Recordset
    Dim frmDrug As frmDrugInfo
    
    glngModule = lngModule
            
    '1.�жϵ�ǰҵ���Ƿ���Ҫ�ϴ�����
    If GetJudge_IsNeedUpload(lngModule, bytType) = False Then Exit Function
        
    '2.ȷ���豸ID�����ܻ��ж��
        
    '3.�ϴ����ݴ�����������ڶ���豸��ѭ������
    Select Case bytType
        Case 0  'ҩƷ������Ϣ�ϴ�
            Set frmDrug = New frmDrugInfo
            frmDrug.Show vbModal, gfrmOwner
            Set frmDrug = Nothing
            
            Upload = True
            Exit Function
        Case 1  '������ϸ�ϴ�
            '��ȡHIS���ݣ����ϴ�
            mdlProcessData.SetUpload bytType, varKey
        Case 2  '��ҩ��Ϣ�ϴ�(������Ҫ��Ϣ)
            '
        Case 3  '��ҪҩƷҽ����Ϣ�ϴ�
            '
        Case 4  'ҩƷ�����Ϣ�ϴ�
            '
        Case Else
            '
    End Select
    
    Upload = True
    Exit Function
    
errHandle:
    '
End Function