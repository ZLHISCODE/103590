VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmWinsock 
   ClientHeight    =   825
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   2250
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   825
   ScaleWidth      =   2250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Visible         =   0   'False
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   1320
      Top             =   120
   End
   Begin MSWinsockLib.Winsock winSock 
      Left            =   360
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
End
Attribute VB_Name = "frmWinsock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event DataArrival(ByVal lngNoticeCode As Long, ByVal intChangeType As Integer, ByVal strTableOwner As String, ByVal strTableName As String, ByVal strRowid As String)
Public blnDcnState As Boolean

Private mlngCheckDcnInterval As Long    '���dcn������״̬ʱ����
Private mlngCheck As Long

Private mcolInterval As New Collection  '����֪ͨInterval
Private mcolData As New Collection '����䶯��Ϣ
Private mcolTime As New Collection '����䶯��Ϣ����ʱ��

Public Function StartUdp(ByVal lngPort As Long, Optional ByRef strError As String) As Boolean
    '�����˿�,�ɹ�����True
    On Error Resume Next

    winSock.LocalPort = lngPort
    winSock.Bind

    If Err.Number <> 0 Then
        strError = Err.Description
        Exit Function
    End If
    StartUdp = True
End Function


Private Sub winSock_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim intType As Integer, strOwner As String, strTable As String
    Dim strRowid As String, lngNoticeCode As Long
    Dim lngInterval As Long, arrTmp() As String
    
    On Error GoTo errH
    
    winSock.GetData strData
    If strData = "" Then Exit Sub
    
    '��ȡ���DCN��ʱ����
    If mlngCheckDcnInterval = 0 Then
        mlngCheckDcnInterval = GetCheckInterval
        blnDcnState = GetDcnState
    End If
    
    '��Ϣ��ʽ:   NoticeCode-�䶯����-������-����-Rowid
    arrTmp = Split(strData, "-")
    lngNoticeCode = arrTmp(0): intType = arrTmp(1)
    strOwner = arrTmp(2): strTable = arrTmp(3): strRowid = arrTmp(4)
    
    lngInterval = GetNoticeInterval(lngNoticeCode)
    
    If lngInterval = 0 Then '���IntervalΪ0 ,˵������Ҫ���л���,ֱ���׳��¼�
        RaiseEvent DataArrival(lngNoticeCode, intType, strOwner, strTable, strRowid)
    Else
        If GetValueFromList(mcolInterval, lngNoticeCode) = "" Then  '������û���ҵ�noticecode,����ӵ�������,�ﵽ�趨ʱ������׳��¼�
            mcolData.Add strData, "_" & lngNoticeCode
            mcolInterval.Add lngInterval, "_" & lngNoticeCode
            mcolTime.Add 0, "_" & lngNoticeCode
        End If
    End If
    
    Exit Sub
errH:
    gobjComLib.ErrCenter
End Sub


Private Sub Timer_Timer()
    Dim i As Integer, arrTmp() As String
    Dim intType As Integer, strOwner As String, strTable As String
    Dim strRowid As String, lngNoticeCode As Long
    Dim lngInterval As Long, arrRemove() As Long
    
    '���ݻ���
    ReDim arrRemove(0)
    
    For i = 1 To mcolTime.Count
        mcolTime.Item(i) = mcolTime.Item(i) + 1
        
        If mcolTime.Item(i) >= mcolInterval.Item(i) Then    '�ﵽˢ��ʱ��
            arrTmp = Split(mcolData.Item(i), "-")
            lngNoticeCode = arrTmp(0): intType = arrTmp(1)
            strOwner = arrTmp(2): strTable = arrTmp(3): strRowid = arrTmp(4)
                    
            RaiseEvent DataArrival(lngNoticeCode, intType, strOwner, strTable, strRowid)    '�׳��¼�
            
            ReDim Preserve arrRemove(UBound(arrRemove) + 1)     '���Ѿ��׳��¼���֪ͨ��¼��������
            arrRemove(UBound(arrRemove)) = lngNoticeCode
        End If
    Next
    
    For i = 1 To UBound(arrRemove) 'ѭ������,ɾ�����׳��¼�֪ͨ
        mcolData.Remove "_" & arrRemove(i)
        mcolInterval.Remove "_" & arrRemove(i)
        mcolTime.Remove "_" & arrRemove(i)
    Next
    
    'Dcn״̬���
    If mlngCheckDcnInterval <> 0 Then
        mlngCheck = mlngCheck + 1
        
        If mlngCheck > mlngCheckDcnInterval Then
            blnDcnState = GetDcnState
            mlngCheck = 0
        End If
    End If
End Sub

Private Function GetNoticeInterval(lngNoticeCode As Long) As Long
    '����:����NoticeCode��ȡInterval
    Dim strSql As String, rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = "Select Interval From zltools.ZlnoticeLists where NoticeCode = [1]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡInterval", lngNoticeCode)
    
    If rsTmp.RecordCount > 0 Then
        GetNoticeInterval = Val(rsTmp!Interval & "")
    End If
    
    Exit Function
errH:
    gobjComLib.ErrCenter
End Function

Private Function GetValueFromList(colList As Collection, lngIndex As Long) As String
    '����:����Index�ڼ����л�ȡֵ,���û���ҵ��򷵻ؿ�
    Dim strResult As String
    
    On Error Resume Next
    
    strResult = colList.Item(lngIndex)
    
    If Err.Number = 0 Then
        GetValueFromList = strResult
    End If
End Function


Private Function GetCheckInterval() As Long
    '��ȡ "DCN���ʱ����¼��"
    Dim rsTmp As New ADODB.Recordset
       
    Set rsTmp = GetZLOptions(32)
    If rsTmp.RecordCount > 0 Then GetCheckInterval = Val(rsTmp!����ֵ)
End Function

Private Function GetDcnState() As Boolean
    '���DCN�������Ƿ���������
    Dim strSql As String, rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select 1" & vbNewLine & _
                "From (Select To_Date(����ֵ, 'yyyy-mm-dd hh24:mi:ss') ���ʱ�� From zlOptions Where ������ = 31) A," & vbNewLine & _
                "     (Select ����ֵ ���¼�� From zlOptions Where ������ = 32) B" & vbNewLine & _
                "Where Sysdate < a.���ʱ�� + b.���¼�� / 24 / 60"

    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "DCN���ʱ����¼��")
    
    GetDcnState = rsTmp.RecordCount > 0
    Exit Function
errH:
    gobjComLib.ErrCenter
End Function


