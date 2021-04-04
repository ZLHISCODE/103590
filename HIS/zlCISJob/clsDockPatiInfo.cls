VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDockPatiInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents mfrmDockPatiInfo As frmDockPatiInfo
Attribute mfrmDockPatiInfo.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set mfrmDockPatiInfo = New frmDockPatiInfo
End Sub

Public Function zlGetForm() As Object
    Set zlGetForm = mfrmDockPatiInfo
End Function

Private Sub Class_Terminate()
    On Error Resume Next
    Unload mfrmDockPatiInfo
    Set mfrmDockPatiInfo = Nothing
    err.Clear
End Sub

Public Sub zlRefresh(ByVal lng����ID As Long, ByVal lng�Һ�id As Long, ByVal blnEdit As Boolean, ByVal blnMoved As Boolean, Optional ByRef objMip As Object, Optional ByVal int���� As Integer)
    Call mfrmDockPatiInfo.zlRefresh(lng����ID, lng�Һ�id, False, blnMoved, objMip, int����)
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
    Call mfrmDockPatiInfo.SetFontSize(bytSize)
End Sub

Public Function ShowSentenceSel(frmParent As Object, ByVal lng�����ļ�id As Long, ByVal str�Ա� As String, ByVal str����״�� As String, ByVal strType As String, _
    Optional ByVal strInput As String, Optional ByVal lngInputHwnd As Long, Optional blnCancel As Boolean) As String
'���ܣ��ʾ�ѡ��
    ShowSentenceSel = frmSentenceSel.ShowMe(frmParent, lng�����ļ�id, str�Ա�, str����״��, strType, strInput, lngInputHwnd, blnCancel)
End Function

Public Function ShowImportEPRDemo(frmParent As Object, ByVal lngFileID As Long, ByVal lngPatient As Long, ByVal lngVisit As Long, ByRef rsInfo As Recordset) As Long
'���ܣ����뷶��
    ShowImportEPRDemo = frmImportEPRDemo.ShowMe(frmParent, lngFileID, lngPatient, lngVisit, rsInfo)
End Function

Public Function ShowOutDocterSign(ByRef fParent As Object, ByVal strSource As String, _
    lngPatiID As Long, lngPatiPageID As Long) As cEPRSign
'���ܣ�ǩ��
    Set ShowOutDocterSign = frmOutDocterSign.ShowMe(fParent, strSource, lngPatiID, lngPatiPageID)
End Function