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

Public Sub zlRefresh(ByVal lng病人ID As Long, ByVal lng挂号id As Long, ByVal blnEdit As Boolean, ByVal blnMoved As Boolean, Optional ByRef objMip As Object, Optional ByVal int调用 As Integer)
    Call mfrmDockPatiInfo.zlRefresh(lng病人ID, lng挂号id, False, blnMoved, objMip, int调用)
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
    Call mfrmDockPatiInfo.SetFontSize(bytSize)
End Sub

Public Function ShowSentenceSel(frmParent As Object, ByVal lng病历文件id As Long, ByVal str性别 As String, ByVal str婚姻状况 As String, ByVal strType As String, _
    Optional ByVal strInput As String, Optional ByVal lngInputHwnd As Long, Optional blnCancel As Boolean) As String
'功能：词句选择
    ShowSentenceSel = frmSentenceSel.ShowMe(frmParent, lng病历文件id, str性别, str婚姻状况, strType, strInput, lngInputHwnd, blnCancel)
End Function

Public Function ShowImportEPRDemo(frmParent As Object, ByVal lngFileID As Long, ByVal lngPatient As Long, ByVal lngVisit As Long, ByRef rsInfo As Recordset) As Long
'功能：导入范文
    ShowImportEPRDemo = frmImportEPRDemo.ShowMe(frmParent, lngFileID, lngPatient, lngVisit, rsInfo)
End Function

Public Function ShowOutDocterSign(ByRef fParent As Object, ByVal strSource As String, _
    lngPatiID As Long, lngPatiPageID As Long) As cEPRSign
'功能：签名
    Set ShowOutDocterSign = frmOutDocterSign.ShowMe(fParent, strSource, lngPatiID, lngPatiPageID)
End Function