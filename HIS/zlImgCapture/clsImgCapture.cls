VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImgCapture"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InitImgCapture(cnMain As ADODB.Connection)
    Set gcnOracle = cnMain
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
End Sub

Public Sub ImageCapture(ByVal strPrivs As String, ByVal lngAdviceID As Long, ByVal lngSendNO As Long, frmObject As Object, strNO As String, int��¼���� As Integer, _
           lng����ID, lng����ID, strҽ������ As String, Optional ByVal strImageType As String = "", Optional ByVal strCheckUID As String = "")
    blCompressionStup = False
    blClosefrm = False
    frmImgCapture.ShowMe strPrivs, lngAdviceID, lngSendNO, frmObject, strNO, int��¼����, lng����ID, lng����ID, strҽ������, strImageType, strCheckUID
End Sub

Public Sub ImageScan(ByVal lngAdviceID As Long, ByVal lngSendNO As Long, Optional ByVal strImageType As String = "", _
    Optional ByVal strCheckUID As String = "")
    frmImgScan.ShowMe lngAdviceID, lngSendNO, strImageType, strCheckUID
End Sub
Public Sub UnladImgCapture()
    Unload frmImgCapture
End Sub
