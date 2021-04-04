VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArchiveMedRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mfrmInMedRec As Object
Attribute mfrmInMedRec.VB_VarHelpID = -1
Private mfrmOutMedRec As frmArchiveOutMedRec

Public Sub InitArchiveMedRec(cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal strDBUser As String)
    Set gcnOracle = cnMain
    
    glngSys = lngSys
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    gstrUnitName = GetUnitName
    gstrDBUser = strDBUser
End Sub

Private Sub Class_Terminate()
    If Not mfrmInMedRec Is Nothing Then
        Unload mfrmInMedRec
        Set mfrmInMedRec = Nothing
    End If
    If Not mfrmOutMedRec Is Nothing Then
        Unload mfrmOutMedRec
        Set mfrmOutMedRec = Nothing
    End If
End Sub

Public Function zlGetForm(ByVal intType As Integer) As Object
'功能：获取子业务窗口,由主程序初始化时只调用一次获取
'参数：intType 0-门诊病案查阅 1-住院病案查阅
    Dim intMecStandard As Integer  '病案首页格式 0-卫生部标准，1-四川省标准，2-云南省标准,3-湖南省标准
    
    If intType = 0 Then
        If mfrmOutMedRec Is Nothing Then
            Set mfrmOutMedRec = New frmArchiveOutMedRec
        End If
        Set zlGetForm = mfrmOutMedRec
    Else
        If mfrmInMedRec Is Nothing Then
            '病案首页标准
            intMecStandard = Val(zlDatabase.GetPara("病案首页标准", glngSys, p住院医生站, "0"))
            Select Case intMecStandard
                Case 0 '卫生部标准
                    Set mfrmInMedRec = New frmArchiveInMedRec
                Case 1    '四川省标准
                    Set mfrmInMedRec = New frmArchiveInMedRec_SC
                Case 2    '云南省标准
                    Set mfrmInMedRec = New frmArchiveInMedRec_YN
                Case 3    '湖南省标准
                    Set mfrmInMedRec = New frmArchiveInMedRec_HN
            End Select
        End If
        Set zlGetForm = mfrmInMedRec
    End If
End Function

Public Function zlRefresh(ByVal intType As Integer, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal blnMoved As Boolean) As Boolean
'功能：刷新或清除医嘱清单
'参数：intType 0-门诊病案查阅 1-住院病案查阅
'      lng主页ID :住院首页-主页ID， 门诊首页-挂号ID
'      blnMoved 是否转储
    If intType = 0 Then
        Call mfrmOutMedRec.zlRefresh(lng病人ID, lng主页ID, blnMoved)
    Else
        Call mfrmInMedRec.zlRefresh(lng病人ID, lng主页ID, blnMoved)
    End If
End Function


